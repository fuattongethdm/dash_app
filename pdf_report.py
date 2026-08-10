"""
PDF report generation: builds a printable A3 report of the dashboard for a
given report date (charts + summary + project detail table).

Mirrors the current Dashboard tab in pages/home.py: the same trend, skelp
impact, pareto and bubble sections, rendered statically with matplotlib
instead of Plotly (a static, non-interactive PDF has no use for the
dashboard's checklists/dropdowns, so every line/series that a checklist
would gate is just always drawn).
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
    amount_in_display_unit,
    apply_meter_based_repair_ratios,
    daily_weighted_repair_ratios,
    daily_weighted_repair_ratios_for_type,
    length_in_display_unit,
    repair_amount_trend_data,
    unit_label,
)

_ACCENT = "#2563eb"  # Coil / primary (matches dashboard COLOR_COIL)
_ACCENT2 = "#f97316"  # Incl. skelp / impact / secondary (matches COLOR_SECONDARY)
_PLATE = "#7c3aed"  # Plate (matches dashboard COLOR_PLATE)
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


def _project_x_labels(ax, labels: pd.Series) -> None:
    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels(labels, rotation=45, ha="right", fontsize=7)


def _trend_chart(overall_ratio: pd.DataFrame) -> Image:
    fig, ax = plt.subplots(figsize=(7, 3.6))
    ax.plot(
        overall_ratio["date"],
        overall_ratio["weighted_repair_ratio"] * 100,
        marker="o",
        color=_ACCENT,
        linewidth=2,
        label="Excl. Skelp",
    )
    ax.plot(
        overall_ratio["date"],
        overall_ratio["weighted_repair_ratio_incl_skelp"] * 100,
        marker="o",
        color=_ACCENT2,
        linewidth=2,
        linestyle="--",
        label="Incl. Skelp",
    )
    ax.set_ylabel("Repair Rate (%)")
    ax.legend(fontsize=8, frameon=False)
    _style_axes(ax, "Overall Repair Rate Trend")
    fig.autofmt_xdate()
    fig.tight_layout()
    return _figure_to_image(fig, _HALF_WIDTH)


def _type_trend_chart(master_df: pd.DataFrame, baseline_df: pd.DataFrame | None = None) -> Image:
    production_types = sorted(master_df["production_type"].dropna().unique())
    type_colors = {"Coil": _ACCENT, "Plate": _PLATE}
    overall_ratio = daily_weighted_repair_ratios(master_df, baseline_df)

    fig, ax = plt.subplots(figsize=(7, 3.6))
    for p_type in production_types:
        type_trend = daily_weighted_repair_ratios_for_type(master_df, p_type, baseline_df)
        ax.plot(
            type_trend["date"],
            type_trend["weighted_repair_ratio"] * 100,
            marker="o",
            linewidth=2,
            color=type_colors.get(p_type),
            label=p_type,
        )
    ax.plot(
        overall_ratio["date"],
        overall_ratio["weighted_repair_ratio"] * 100,
        marker="o",
        linewidth=2,
        linestyle="--",
        color=_ACCENT2,
        label="Mix (Coil + Plate)",
    )
    ax.set_ylabel("Repair Rate (%)")
    ax.legend(fontsize=8, frameon=False)
    _style_axes(ax, "Repair Rate Trend by Production Type")
    fig.autofmt_xdate()
    fig.tight_layout()
    return _figure_to_image(fig, _HALF_WIDTH)


def _amount_chart(daily_amount: pd.DataFrame, unit: str) -> Image:
    fig, ax = plt.subplots(figsize=(14, 3.6))
    ax.bar(daily_amount["date"], daily_amount["daily_repair_amount_display"], color=_ACCENT)
    ax.set_ylabel(f"Amount ({unit})")
    _style_axes(ax, "Daily Repair Amount")
    fig.autofmt_xdate()
    fig.tight_layout()
    return _figure_to_image(fig, _USABLE_WIDTH)


def _worst_projects_chart(latest_df: pd.DataFrame) -> Image:
    worst = latest_df.sort_values("repair_ratio", ascending=False).head(10).iloc[::-1]
    fig, ax = plt.subplots(figsize=(14, 3.6))
    ax.barh(worst["project_no"], worst["repair_ratio"], color=_ACCENT)
    ax.set_xlabel("Repair Ratio")
    _style_axes(ax, "Top 10 Projects by Repair Ratio (Latest Day)")
    fig.tight_layout()
    return _figure_to_image(fig, _USABLE_WIDTH)


def _skelp_impact_charts(latest_df: pd.DataFrame, unit: str) -> tuple[Image, Image]:
    """Same top-10 projects (ranked by ratio impact) in both charts, so
    they stay directly comparable — mirrors the dashboard's skelp-impact
    pair (pages/home.py, render_dashboard)."""
    ratio_df = latest_df[["project_label", "repair_ratio", "repair_ratio_incl_skelp"]].copy()
    ratio_df["skelp_impact"] = (ratio_df["repair_ratio_incl_skelp"] - ratio_df["repair_ratio"]).round(6)
    ratio_top = ratio_df.sort_values("skelp_impact", ascending=False).head(10).reset_index(drop=True)

    amount_df = latest_df[["project_label", "total_repair_amount", "total_repair_amount_incl_skelp"]].copy()
    amount_df["skelp_impact_amount"] = (
        amount_df["total_repair_amount_incl_skelp"] - amount_df["total_repair_amount"]
    ).round(4)
    amount_top = ratio_top[["project_label"]].merge(amount_df, on="project_label", how="left")

    fig1, ax1 = plt.subplots(figsize=(7, 4.2))
    x = range(len(ratio_top))
    ax1.bar(x, ratio_top["repair_ratio"] * 100, color=_ACCENT, label="Repair Ratio")
    ax1.bar(x, ratio_top["skelp_impact"] * 100, bottom=ratio_top["repair_ratio"] * 100, color=_ACCENT2, label="Skelp Impact")
    _project_x_labels(ax1, ratio_top["project_label"])
    ax1.set_ylabel("Repair Ratio (%)")
    ax1.legend(fontsize=8, frameon=False)
    _style_axes(ax1, "Skelp Impact on Repair Ratio (Top 10, Latest Day)")
    fig1.tight_layout()
    ratio_img = _figure_to_image(fig1, _HALF_WIDTH)

    fig2, ax2 = plt.subplots(figsize=(7, 4.2))
    ax2.bar(x, amount_top["total_repair_amount"], color=_ACCENT, label=f"Repair Amount ({unit})")
    ax2.bar(
        x,
        amount_top["skelp_impact_amount"],
        bottom=amount_top["total_repair_amount"],
        color=_ACCENT2,
        label=f"Skelp Impact ({unit})",
    )
    _project_x_labels(ax2, amount_top["project_label"])
    ax2.set_ylabel(f"Repair Amount ({unit})")
    ax2.legend(fontsize=8, frameon=False)
    _style_axes(ax2, "Skelp Impact on Repair Amount (Same Top 10 as Ratio Impact)")
    fig2.tight_layout()
    amount_img = _figure_to_image(fig2, _HALF_WIDTH)

    return ratio_img, amount_img


def _pareto_chart(latest_df: pd.DataFrame, value_col: str, y_label: str, title: str, as_pct: bool) -> Image:
    ordered = latest_df[["project_label", value_col]].sort_values(value_col, ascending=False).reset_index(drop=True)
    ordered["cumulative_pct"] = ordered[value_col].cumsum() / ordered[value_col].sum() * 100
    values = ordered[value_col] * 100 if as_pct else ordered[value_col]

    fig, ax1 = plt.subplots(figsize=(14, 4.2))
    x = range(len(ordered))
    ax1.bar(x, values, color=_ACCENT)
    _project_x_labels(ax1, ordered["project_label"])
    ax1.set_ylabel(y_label)

    ax2 = ax1.twinx()
    ax2.plot(x, ordered["cumulative_pct"], color=_ACCENT2, marker="o", linewidth=2)
    ax2.set_ylabel("Cumulative %")
    ax2.set_ylim(0, 105)
    ax2.spines["top"].set_visible(False)

    _style_axes(ax1, title)
    fig.tight_layout()
    return _figure_to_image(fig, _USABLE_WIDTH)


def _bubble_chart(latest_df: pd.DataFrame, x_col: str, x_label: str, title: str) -> Image:
    sizes = latest_df["total_repair_amount"]
    max_size = sizes.max() or 1
    # Area-based marker sizing (matplotlib's `s` is area in points^2),
    # scaled so the biggest bubble stays readable without swamping the rest.
    scaled_sizes = (sizes / max_size) * 900 + 30

    fig, ax = plt.subplots(figsize=(14, 5))
    sc = ax.scatter(
        latest_df[x_col],
        latest_df["repair_ratio_pct"],
        s=scaled_sizes,
        c=latest_df["repair_ratio_pct"],
        cmap="YlOrRd",
        alpha=0.55,
        edgecolors="black",
        linewidths=1.2,
    )
    cbar = fig.colorbar(sc, ax=ax)
    cbar.set_label("Repair Ratio (%)")
    ax.set_xlabel(x_label)
    ax.set_ylabel("Spiral Repair Ratio (%)")
    _style_axes(ax, title)
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


def _side_by_side(left: Image, right: Image) -> Table:
    row = Table([[left, right]], colWidths=[_HALF_WIDTH, _HALF_WIDTH])
    row.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    return row


def build_pdf_report(master_df: pd.DataFrame, baseline_df: pd.DataFrame, selected_date, display_unit: str = "m") -> bytes:
    """Build an A3-landscape PDF report for the given report date.

    `master_df` is the same frame used to render the dashboard (see
    pages/home.py:render_dashboard). `display_unit` ("m" or "ft") mirrors
    the dashboard's unit toggle at the time the PDF was requested.
    """
    u = unit_label(display_unit)
    latest_df = master_df[master_df["date"] == selected_date].copy()
    latest_df = apply_meter_based_repair_ratios(latest_df)
    # Same fix as the dashboard: project_no alone repeats across dimensions,
    # so every per-project chart below keys off project_label instead.
    latest_df["project_label"] = latest_df["project_no"].astype(str) + " (" + latest_df["dimensions"].astype(str) + ")"
    latest_df["repair_ratio_pct"] = (latest_df["repair_ratio"] * 100).round(4)

    # Ratios above are computed from the raw feet/meter columns (the
    # dashboard standard) — only convert to the display unit afterwards, so
    # every chart below shows/scales consistently with the unit toggle.
    latest_df["total_repair_amount"] = amount_in_display_unit(latest_df["total_repair_amount"], display_unit)
    latest_df["total_repair_amount_incl_skelp"] = amount_in_display_unit(
        latest_df["total_repair_amount_incl_skelp"], display_unit
    )
    latest_df["project_total_pipe_length"] = length_in_display_unit(latest_df["project_total_pipe_length"], display_unit)
    latest_df["repaired_spiral_length"] = length_in_display_unit(latest_df["repaired_spiral_length"], display_unit)

    overall_ratio = daily_weighted_repair_ratios(master_df, baseline_df)
    current_ratio = (
        overall_ratio.loc[overall_ratio["date"] == selected_date, "weighted_repair_ratio"].iloc[0]
        if not overall_ratio.empty
        else 0
    )
    daily_amount = repair_amount_trend_data(master_df, display_unit=display_unit)

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
    section_style = styles["Heading2"]

    skelp_ratio_img, skelp_amount_img = _skelp_impact_charts(latest_df, u)

    story = [
        Paragraph("Daily Repair Rate Report", title_style),
        Paragraph(
            f"Report Date: {pd.to_datetime(selected_date).strftime('%d.%m.%Y')}  |  "
            f"Active Projects: {latest_df['project_no'].nunique()}  |  "
            f"Overall Repair Rate: {current_ratio * 100:.2f}%",
            subtitle_style,
        ),
        Spacer(1, 0.4 * cm),
        _side_by_side(_trend_chart(overall_ratio), _type_trend_chart(master_df, baseline_df)),
        Spacer(1, 0.3 * cm),
        _worst_projects_chart(latest_df),
        Spacer(1, 0.3 * cm),
        _side_by_side(skelp_ratio_img, skelp_amount_img),
        Spacer(1, 0.3 * cm),
        _pareto_chart(latest_df, "repair_ratio", "Repair Ratio (%)", "Repair Ratio Pareto (Latest Day)", as_pct=True),
        Spacer(1, 0.3 * cm),
        _pareto_chart(
            latest_df, "total_repair_amount", f"Repair Amount ({u})", f"Repair Amount Pareto ({u}, Latest Day)", as_pct=False
        ),
        Spacer(1, 0.3 * cm),
        _bubble_chart(
            latest_df,
            "project_total_pipe_length",
            f"Total Production Pipe Length ({u})",
            "Repair Ratio vs. Production Volume by Project (Latest Day)",
        ),
        Spacer(1, 0.3 * cm),
        _bubble_chart(
            latest_df,
            "repaired_spiral_length",
            f"Total Spiral Length ({u})",
            "Repair Ratio vs. Spiral Length by Project (Latest Day)",
        ),
        Spacer(1, 0.3 * cm),
        _amount_chart(daily_amount, u),
        Spacer(1, 0.5 * cm),
        Paragraph("Latest Day — Project Details", section_style),
        Spacer(1, 0.2 * cm),
        _detail_table(latest_df),
    ]
    doc.build(story)
    return buffer.getvalue()
