"""
PDF report generation: builds a printable A3 report of the dashboard for a
given report date (charts + summary + project detail table).
"""

from __future__ import annotations

import io

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from calculations import (
    apply_meter_based_repair_ratios,
    daily_weighted_repair_ratios,
    repair_amount_trend_data,
)

_ACCENT = "#2563eb"
_GRID = "#e2e8f0"
_ROW_ALT = "#f4f6f8"

# A3 landscape usable width with 1.5cm margins on each side.
_PAGE_MARGIN = 1.5 * cm
_USABLE_WIDTH = landscape(A3)[0] - 2 * _PAGE_MARGIN
_HALF_WIDTH = (_USABLE_WIDTH - 0.5 * cm) / 2


def _figure_to_image(fig, width: float) -> Image:
    """Render a matplotlib figure to a reportlab Image at the given width,
    computing height from the actual saved PNG's pixel aspect ratio (not
    the nominal figsize) so bbox_inches="tight" cropping can never distort
    the image the way a hardcoded width/height pair would.
    """
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=150, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    aspect = PILImage.open(buf).height / PILImage.open(buf).width
    buf.seek(0)
    return Image(buf, width=width, height=width * aspect)


def _style_axes(ax, title: str) -> None:
    ax.set_title(title, fontsize=11, fontweight="bold")
    ax.grid(alpha=0.3)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)


def _trend_chart(overall_ratio: pd.DataFrame) -> Image:
    fig, ax = plt.subplots(figsize=(7, 3.6))
    ax.plot(
        overall_ratio["date"],
        overall_ratio["weighted_repair_ratio"] * 100,
        marker="o",
        color=_ACCENT,
        linewidth=2,
    )
    ax.set_ylabel("Repair Rate (%)")
    _style_axes(ax, "Overall Repair Rate Trend")
    fig.autofmt_xdate()
    fig.tight_layout()
    return _figure_to_image(fig, _HALF_WIDTH)


def _amount_chart(daily_amount: pd.DataFrame) -> Image:
    fig, ax = plt.subplots(figsize=(7, 3.6))
    ax.bar(daily_amount["date"], daily_amount["daily_repair_amount_display"], color=_ACCENT)
    ax.set_ylabel("Amount (m)")
    _style_axes(ax, "Daily Repair Amount")
    fig.autofmt_xdate()
    fig.tight_layout()
    return _figure_to_image(fig, _HALF_WIDTH)


def _worst_projects_chart(latest_df: pd.DataFrame) -> Image:
    worst = latest_df.sort_values("repair_ratio", ascending=False).head(10).iloc[::-1]
    fig, ax = plt.subplots(figsize=(14, 3.6))
    ax.barh(worst["project_no"], worst["repair_ratio"], color=_ACCENT)
    ax.set_xlabel("Repair Ratio")
    _style_axes(ax, "Top 10 Projects by Repair Ratio (Latest Day)")
    fig.tight_layout()
    return _figure_to_image(fig, _USABLE_WIDTH)


def _detail_table(latest_df: pd.DataFrame) -> Table:
    columns = ["project_no", "dimensions", "production_type", "qty", "project_status", "repair_ratio"]
    header = ["Project", "Dimensions", "Type", "Qty", "Status", "Repair Ratio"]
    data = [header]
    for _, row in latest_df[columns].sort_values("repair_ratio", ascending=False).iterrows():
        data.append(
            [
                str(row["project_no"]),
                str(row["dimensions"]),
                str(row["production_type"]),
                f"{row['qty']:,.0f}",
                str(row["project_status"]),
                f"{row['repair_ratio']:.2%}",
            ]
        )

    # Spread the six columns across the full page width instead of
    # shrink-wrapping to content, so the table matches the charts above it.
    col_fractions = [0.30, 0.24, 0.12, 0.10, 0.14, 0.10]
    col_widths = [_USABLE_WIDTH * f for f in col_fractions]

    table = Table(data, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(_ACCENT)),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor(_GRID)),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor(_ROW_ALT)]),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 5),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
            ]
        )
    )
    return table


def build_pdf_report(master_df: pd.DataFrame, baseline_df: pd.DataFrame, selected_date) -> bytes:
    """Build an A3-landscape PDF report for the given report date.

    `master_df` and `baseline_df` are the same frames used to render the
    dashboard (see pages/home.py:render_dashboard).
    """
    latest_df = master_df[master_df["date"] == selected_date].copy()
    latest_df = apply_meter_based_repair_ratios(latest_df)

    overall_ratio = daily_weighted_repair_ratios(master_df, baseline_df)
    current_ratio = (
        overall_ratio.loc[overall_ratio["date"] == selected_date, "weighted_repair_ratio"].iloc[0]
        if not overall_ratio.empty
        else 0
    )
    daily_amount = repair_amount_trend_data(master_df, display_unit="m")

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A3),
        leftMargin=_PAGE_MARGIN,
        rightMargin=_PAGE_MARGIN,
        topMargin=_PAGE_MARGIN,
        bottomMargin=_PAGE_MARGIN,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("ReportTitle", parent=styles["Title"], fontSize=20, spaceAfter=4)
    subtitle_style = ParagraphStyle("ReportSubtitle", parent=styles["Normal"], fontSize=11, textColor=colors.HexColor("#64748b"))

    # Trend + Amount charts side by side, sharing the full page width.
    charts_row = Table(
        [[_trend_chart(overall_ratio), _amount_chart(daily_amount)]],
        colWidths=[_HALF_WIDTH, _HALF_WIDTH],
    )
    charts_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))

    story = [
        Paragraph("Daily Repair Rate Report", title_style),
        Paragraph(
            f"Report Date: {pd.to_datetime(selected_date).strftime('%d.%m.%Y')}  |  "
            f"Active Projects: {latest_df['project_no'].nunique()}  |  "
            f"Overall Repair Rate: {current_ratio * 100:.2f}%",
            subtitle_style,
        ),
        Spacer(1, 0.4 * cm),
        charts_row,
        Spacer(1, 0.3 * cm),
        _worst_projects_chart(latest_df),
        Spacer(1, 0.5 * cm),
        Paragraph("Latest Day — Project Details", styles["Heading2"]),
        Spacer(1, 0.2 * cm),
        _detail_table(latest_df),
    ]
    doc.build(story)
    return buffer.getvalue()
