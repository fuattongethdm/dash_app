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
import re

import matplotlib

matplotlib.use("Agg")
import matplotlib.dates as mdates
import matplotlib.pyplot as plt
import pandas as pd
from PIL import Image as PILImage
from reportlab.lib import colors
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from calculations import (
    METERS_PER_FOOT,
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

# Keep these two in sync with the matching constants in pages/home.py — the
# PDF is a static mirror of the dashboard, rendered with a different
# plotting library, so the values are duplicated rather than shared.
TREND_WINDOW_DAYS = 20
BUBBLE_LABEL_TOP_N = 5


def _strip_quarter_prefix(project_no) -> str:
    """'2026Q-10-108JT' -> '10-108JT' — same "20XXQ-" strip the dashboard's
    repair-ratio charts use, so the PDF axis labels match."""
    return re.sub(r"^\d{4}Q-", "", str(project_no))


def _add_bar_value_annotations(ax, positions, values, display_values, min_fraction: float = 0.15) -> None:
    """Write ``display_values`` (pipe quantity) inside each vertical bar at
    its mid-height, skipping any bar under ``min_fraction`` of the chart's
    max value — the same "won't fit without overflowing" heuristic used in
    pages/home.py."""
    values = list(values)
    max_value = max(values) if values else 0
    if not max_value:
        return
    for position, value, display_value in zip(positions, values, display_values):
        if value < max_value * min_fraction:
            continue
        ax.text(position, value / 2, str(display_value), ha="center", va="center", color="white", fontsize=8)


def _add_line_value_labels(ax, x_values, y_values, color, va: str = "bottom", fontsize: int = 8) -> None:
    """Percent value next to each point on a trend line — matches the
    dashboard's always-visible line labels (pages/home.py). ``va="bottom"``
    puts the text above the point, ``"top"`` puts it below, so two series
    on the same chart don't collide."""
    offset = 6 if va == "bottom" else -10
    for x, y in zip(x_values, y_values):
        ax.annotate(
            f"{y:.2f}%",
            (x, y),
            xytext=(0, offset),
            textcoords="offset points",
            fontsize=fontsize,
            color=color,
            ha="center",
            va=va,
        )


def _figure_to_image(fig, width: float) -> Image:
    """Render a matplotlib figure to a reportlab Image at the given width,
    computing height from the actual saved PNG's pixel aspect ratio (not
    the nominal figsize) so bbox_inches="tight" cropping can never distort
    the image the way a hardcoded width/height pair would.
    """
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=170, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    aspect = PILImage.open(buf).height / PILImage.open(buf).width
    buf.seek(0)
    return Image(buf, width=width, height=width * aspect)


def _style_axes(ax, title: str) -> None:
    ax.set_title(title, fontsize=13, fontweight="bold")
    ax.set_ylabel(ax.get_ylabel(), fontsize=10)
    ax.set_xlabel(ax.get_xlabel(), fontsize=10)
    ax.tick_params(axis="both", labelsize=9)
    ax.grid(alpha=0.3)
    for spine in ("top", "right"):
        ax.spines[spine].set_visible(False)


def _project_x_labels(ax, labels: pd.Series, rotation: int = 45) -> None:
    ax.set_xticks(range(len(labels)))
    ax.set_xticklabels(labels, rotation=rotation, ha="right", fontsize=8)


def _date_axis(ax) -> None:
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m.%y"))
    for label in ax.get_xticklabels():
        label.set_rotation(45)
        label.set_ha("right")
        label.set_fontsize(9)


def _trend_chart(overall_ratio: pd.DataFrame) -> Image:
    window_start = overall_ratio["date"].max() - pd.Timedelta(days=TREND_WINDOW_DAYS)
    overall_ratio = overall_ratio[overall_ratio["date"] >= window_start].copy()
    # Round to the same 2-decimal-percent precision as the on-chart label
    # (round(ratio, 4) == round(ratio * 100, 2) / 100) so points that display
    # the same value also plot at the same height.
    overall_ratio["weighted_repair_ratio"] = overall_ratio["weighted_repair_ratio"].round(4)
    overall_ratio["weighted_repair_ratio_incl_skelp"] = overall_ratio["weighted_repair_ratio_incl_skelp"].round(4)

    fig, ax = plt.subplots(figsize=(14, 4))
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
    ax.legend(fontsize=9, frameon=False)
    _style_axes(ax, "Coil and Plate Repair Rate Trend")
    _date_axis(ax)
    _add_line_value_labels(ax, overall_ratio["date"], overall_ratio["weighted_repair_ratio"] * 100, _ACCENT, va="bottom")
    _add_line_value_labels(
        ax, overall_ratio["date"], overall_ratio["weighted_repair_ratio_incl_skelp"] * 100, _ACCENT2, va="top"
    )
    fig.tight_layout()
    return _figure_to_image(fig, _USABLE_WIDTH)


def _type_trend_chart(master_df: pd.DataFrame, baseline_df: pd.DataFrame | None = None) -> Image:
    window_start = master_df["date"].max() - pd.Timedelta(days=TREND_WINDOW_DAYS)
    master_df = master_df[master_df["date"] >= window_start]

    production_types = sorted(master_df["production_type"].dropna().unique())
    type_colors = {"Coil": _ACCENT, "Plate": _PLATE}
    overall_ratio = daily_weighted_repair_ratios(master_df, baseline_df).copy()
    overall_ratio["weighted_repair_ratio"] = overall_ratio["weighted_repair_ratio"].round(4)

    fig, ax = plt.subplots(figsize=(14, 4))
    va_cycle = ["bottom", "top"]
    for i, p_type in enumerate(production_types):
        type_trend = daily_weighted_repair_ratios_for_type(master_df, p_type, baseline_df).copy()
        type_trend["weighted_repair_ratio"] = type_trend["weighted_repair_ratio"].round(4)
        color = type_colors.get(p_type)
        ax.plot(
            type_trend["date"],
            type_trend["weighted_repair_ratio"] * 100,
            marker="o",
            linewidth=2,
            color=color,
            label=p_type,
        )
        _add_line_value_labels(
            ax, type_trend["date"], type_trend["weighted_repair_ratio"] * 100, color, va=va_cycle[i % 2]
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
    _add_line_value_labels(
        ax, overall_ratio["date"], overall_ratio["weighted_repair_ratio"] * 100, _ACCENT2,
        va=va_cycle[len(production_types) % 2],
    )
    ax.set_ylabel("Repair Rate (%)")
    ax.legend(fontsize=9, frameon=False)
    _style_axes(ax, "Repair Rate Trend by Production Type")
    _date_axis(ax)
    fig.tight_layout()
    return _figure_to_image(fig, _USABLE_WIDTH)


def _amount_chart(daily_amount: pd.DataFrame, unit: str) -> Image:
    fig, ax = plt.subplots(figsize=(14, 3.6))
    ax.bar(daily_amount["date"], daily_amount["daily_repair_amount_display"], color=_ACCENT)
    ax.set_ylabel(f"Amount ({unit})")
    _style_axes(ax, f"Daily Repair Amount ({unit})")
    _date_axis(ax)
    fig.tight_layout()
    return _figure_to_image(fig, _USABLE_WIDTH)


def _worst_projects_chart(latest_df: pd.DataFrame) -> Image:
    worst = latest_df.sort_values("repair_ratio", ascending=False).head(10).iloc[::-1].reset_index(drop=True)
    labels = worst["project_no"].map(_strip_quarter_prefix) + " (" + worst["dimensions"].astype(str) + ")"
    fig, ax = plt.subplots(figsize=(14, 3.6))
    ax.barh(labels, worst["repair_ratio"], color=_ACCENT)
    ax.set_xlabel("Repair Ratio")
    _style_axes(ax, "Top 10 Projects by Repair Ratio (Latest Day)")
    fig.tight_layout()
    max_ratio = worst["repair_ratio"].max()
    label_offset = max_ratio * 0.015 if max_ratio else 0
    for y, (value, qty) in enumerate(zip(worst["repair_ratio"], worst["qty"])):
        if value >= max_ratio * 0.15:
            ax.text(value / 2, y, str(int(qty)), ha="center", va="center", color="white", fontsize=8)
        ax.text(value + label_offset, y, f"{value:.2%}", ha="left", va="center", fontsize=8)
    ax.set_xlim(right=max_ratio * 1.1)
    return _figure_to_image(fig, _USABLE_WIDTH)


def _skelp_impact_charts(latest_df: pd.DataFrame, unit: str) -> tuple[Image, Image]:
    """Ratio and amount charts each rank their own top 5 by their own base
    metric (repair_ratio / total_repair_amount) — mirrors the dashboard's
    skelp-impact pair (pages/home.py, render_dashboard). The ratio chart
    drops the "20XXQ-" project prefix and shows pipe quantity inside the
    bar, same as the dashboard's repair-ratio charts; the amount chart
    keeps the plain project label."""
    ratio_df = latest_df[["project_label", "project_no", "dimensions", "qty", "repair_ratio", "repair_ratio_incl_skelp"]].copy()
    ratio_df["skelp_impact"] = (ratio_df["repair_ratio_incl_skelp"] - ratio_df["repair_ratio"]).round(6)
    ratio_top = ratio_df.sort_values("repair_ratio", ascending=False).head(5).reset_index(drop=True)
    ratio_labels = ratio_top["project_no"].map(_strip_quarter_prefix) + " (" + ratio_top["dimensions"].astype(str) + ")"

    amount_df = latest_df[
        ["project_no", "dimensions", "qty", "total_repair_amount", "total_repair_amount_incl_skelp"]
    ].copy()
    amount_df["skelp_impact_amount"] = (
        amount_df["total_repair_amount_incl_skelp"] - amount_df["total_repair_amount"]
    ).round(4)
    amount_top = amount_df.sort_values("total_repair_amount", ascending=False).head(5).reset_index(drop=True)
    amount_labels = amount_top["project_no"].map(_strip_quarter_prefix) + " (" + amount_top["dimensions"].astype(str) + ")"

    fig1, ax1 = plt.subplots(figsize=(7, 4.2))
    x = range(len(ratio_top))
    ratio_pct = ratio_top["repair_ratio"] * 100
    impact_pct = ratio_top["skelp_impact"] * 100
    ax1.bar(x, ratio_pct, color=_ACCENT, label="Repair Ratio")
    ax1.bar(x, impact_pct, bottom=ratio_pct, color=_ACCENT2, label="Skelp Impact")
    _project_x_labels(ax1, ratio_labels, rotation=25)
    ax1.set_ylabel("Repair Ratio (%)")
    ax1.legend(fontsize=9, frameon=False)
    _style_axes(ax1, "Skelp-End Weld Impact on Repair Ratio (Top 5, Latest Day)")
    fig1.tight_layout()
    # Ratio % and qty share the base segment's center — combined into one
    # two-line label instead of a separate overlapping annotation, same fix
    # as the dashboard's skelp_fig (pages/home.py). A segment too small to
    # fit its label gets it written above the bar instead of omitted.
    max_stack = (ratio_pct + impact_pct).max() if len(ratio_pct) else 0
    for xi, r, imp, qty in zip(x, ratio_pct, impact_pct, ratio_top["qty"]):
        base_fits = r >= max_stack * 0.1
        impact_fits = imp >= max_stack * 0.08
        if base_fits:
            ax1.text(xi, r / 2, f"{r:.2f}%\nQty {qty}", ha="center", va="center", color="white", fontsize=7.5)
        if impact_fits:
            ax1.text(xi, r + imp / 2, f"+{imp:.2f}%", ha="center", va="center", color="white", fontsize=8)
        # Impact line first, repair-ratio line last, so the text block reads
        # top-to-bottom in the same order the bar stacks bottom-to-top.
        outside_lines = []
        if not impact_fits:
            outside_lines.append(f"+{imp:.2f}%")
        if not base_fits:
            outside_lines.append(f"{r:.2f}% (Qty {qty})")
        if outside_lines:
            ax1.text(
                xi, r + imp + max_stack * 0.02, "\n".join(outside_lines),
                ha="center", va="bottom", color="#1e293b", fontsize=7,
            )
    ratio_img = _figure_to_image(fig1, _HALF_WIDTH)

    fig2, ax2 = plt.subplots(figsize=(7, 4.2))
    x2 = range(len(amount_top))
    ax2.bar(x2, amount_top["total_repair_amount"], color=_ACCENT, label=f"Repair Amount ({unit})")
    ax2.bar(
        x2,
        amount_top["skelp_impact_amount"],
        bottom=amount_top["total_repair_amount"],
        color=_ACCENT2,
        label=f"Skelp Impact ({unit})",
    )
    _project_x_labels(ax2, amount_labels, rotation=25)
    ax2.set_ylabel(f"Repair Amount ({unit})")
    ax2.legend(fontsize=9, frameon=False)
    _style_axes(ax2, f"Skelp-End Weld Impact on Repair Amount ({unit}, Top 5, Latest Day)")
    fig2.tight_layout()
    # Same base/impact + outside-if-too-small treatment as the ratio chart
    # above, but on the raw amount values instead of percentages.
    amount_max_stack = (
        (amount_top["total_repair_amount"] + amount_top["skelp_impact_amount"]).max() if len(amount_top) else 0
    )
    for xi, v, imp, qty in zip(x2, amount_top["total_repair_amount"], amount_top["skelp_impact_amount"], amount_top["qty"]):
        base_fits = amount_max_stack and v >= amount_max_stack * 0.1
        impact_fits = amount_max_stack and imp >= amount_max_stack * 0.08
        if base_fits:
            ax2.text(xi, v / 2, f"{v:.2f}\nQty {qty}", ha="center", va="center", color="white", fontsize=7.5)
        if impact_fits:
            ax2.text(xi, v + imp / 2, f"+{imp:.2f}", ha="center", va="center", color="white", fontsize=8)
        # Impact line first, repair-amount line last — see the ratio chart above.
        outside_lines = []
        if not impact_fits:
            outside_lines.append(f"+{imp:.2f}")
        if not base_fits:
            outside_lines.append(f"{v:.2f} (Qty {qty})")
        if outside_lines:
            ax2.text(
                xi, v + imp + amount_max_stack * 0.02, "\n".join(outside_lines),
                ha="center", va="bottom", color="#1e293b", fontsize=7,
            )
    amount_img = _figure_to_image(fig2, _HALF_WIDTH)

    return ratio_img, amount_img


def _pareto_chart(latest_df: pd.DataFrame, value_col: str, y_label: str, title: str, as_pct: bool) -> Image:
    ordered = latest_df[["project_no", "dimensions", "qty", value_col]].sort_values(
        value_col, ascending=False
    ).reset_index(drop=True)
    ordered["cumulative_pct"] = ordered[value_col].cumsum() / ordered[value_col].sum() * 100
    values = ordered[value_col] * 100 if as_pct else ordered[value_col]
    labels = ordered["project_no"].map(_strip_quarter_prefix) + " (" + ordered["dimensions"].astype(str) + ")"

    fig, ax1 = plt.subplots(figsize=(14, 4.2))
    x = range(len(ordered))
    ax1.bar(x, values, color=_ACCENT)
    _project_x_labels(ax1, labels)
    ax1.set_ylabel(y_label)

    ax2 = ax1.twinx()
    ax2.plot(x, ordered["cumulative_pct"], color=_ACCENT2, marker="o", linewidth=2)
    ax2.set_ylabel("Cumulative %")
    ax2.set_ylim(0, 105)
    ax2.spines["top"].set_visible(False)
    ax2.tick_params(axis="y", labelsize=9)
    for xi, cum in zip(x, ordered["cumulative_pct"]):
        ax2.annotate(
            f"{cum:.2f}%", (xi, cum), xytext=(0, 6), textcoords="offset points",
            fontsize=7, color=_ACCENT2, ha="center",
        )

    _style_axes(ax1, title)
    fig.tight_layout()
    _add_bar_value_annotations(ax1, x, values, ordered["qty"])
    if as_pct:
        headroom = values.max() * 0.02 if len(values) and values.max() else 0.05
        for xi, value in zip(x, values):
            ax1.text(xi, value + headroom, f"{value:.2f}%", ha="center", fontsize=8)
    return _figure_to_image(fig, _USABLE_WIDTH)


def _dimension_ratio_chart(master_df: pd.DataFrame, selected_date) -> Image:
    """Weighted repair ratio per dimension, top 15 by total repair amount —
    mirrors the dashboard's "Weighted Repair Ratio by Dimension" chart
    (pages/home.py, render_dashboard). Computed from raw feet/meter columns
    (a ratio needs no display-unit conversion)."""
    day_df = master_df[master_df["date"] == selected_date]
    dim_df = (
        day_df.groupby("dimensions", as_index=False)
        .agg(total_repair_amount=("total_repair_amount", "sum"), repaired_spiral_length=("repaired_spiral_length", "sum"))
        .sort_values("total_repair_amount", ascending=False)
        .head(15)
    )
    denom = dim_df["repaired_spiral_length"] * METERS_PER_FOOT
    dim_df["weighted_ratio"] = (dim_df["total_repair_amount"] / denom.where(denom != 0)).fillna(0)
    dim_df = dim_df.sort_values("weighted_ratio", ascending=False).reset_index(drop=True)

    fig, ax = plt.subplots(figsize=(14, 4.2))
    x = range(len(dim_df))
    values = dim_df["weighted_ratio"] * 100
    ax.bar(x, values, color=_ACCENT)
    ax.set_xticks(list(x))
    ax.set_xticklabels(dim_df["dimensions"], rotation=45, ha="right", fontsize=8)
    ax.set_ylabel("Weighted Repair Ratio (%)")
    _style_axes(ax, "Weighted Repair Ratio by Dimension (Latest Day)")
    fig.tight_layout()
    headroom = values.max() * 0.03 if values.max() else 0.1
    for xi, value in zip(x, values):
        ax.text(xi, value + headroom, f"{value:.2f}%", ha="center", fontsize=8)
    return _figure_to_image(fig, _USABLE_WIDTH)


def _bubble_chart(
    latest_df: pd.DataFrame, x_col: str, x_label: str, title: str, label_col: str = "project_label_clean"
) -> Image:
    """label_col is "project_label_clean" (no "20XXQ-" prefix) for the
    by-project variants, or "dimensions" for the by-dimension variant."""
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

    worst = latest_df.nlargest(BUBBLE_LABEL_TOP_N, "repair_ratio_pct")
    for _, row in worst.iterrows():
        ax.annotate(
            str(row[label_col]),
            (row[x_col], row["repair_ratio_pct"]),
            fontsize=8,
            xytext=(0, 8),
            textcoords="offset points",
            ha="center",
        )

    fig.tight_layout()
    return _figure_to_image(fig, _USABLE_WIDTH)


def _dimension_bubble_frame(master_df: pd.DataFrame, selected_date, display_unit: str) -> pd.DataFrame:
    """Aggregate the given day's rows by dimensions for the by-dimension
    bubble variant — mirrors render_bubble_chart's "dimension_spiral"
    branch in pages/home.py. The ratio is computed from the raw feet/meter
    columns first (before any display-unit conversion), then the summed
    length/amount are converted for display, same order as the dashboard."""
    day_df = master_df[master_df["date"] == selected_date]
    dim_df = day_df.groupby("dimensions", as_index=False).agg(
        repaired_spiral_length=("repaired_spiral_length", "sum"),
        total_repair_amount=("total_repair_amount", "sum"),
    )
    denom = dim_df["repaired_spiral_length"] * METERS_PER_FOOT
    dim_df["repair_ratio_pct"] = ((dim_df["total_repair_amount"] / denom.where(denom != 0)).fillna(0) * 100).round(4)
    dim_df["repaired_spiral_length"] = length_in_display_unit(dim_df["repaired_spiral_length"], display_unit)
    dim_df["total_repair_amount"] = amount_in_display_unit(dim_df["total_repair_amount"], display_unit)
    return dim_df


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
                ("FONTSIZE", (0, 0), (-1, -1), 10),
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
    # Prefix-stripped variant for on-chart labels (bubble chart) — same
    # "20XXQ-" strip the pareto/skelp charts already use.
    latest_df["project_label_clean"] = (
        latest_df["project_no"].map(_strip_quarter_prefix) + " (" + latest_df["dimensions"].astype(str) + ")"
    )
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
    subtitle_style = ParagraphStyle("ReportSubtitle", parent=styles["Normal"], fontSize=12, textColor=colors.HexColor("#64748b"))
    section_style = styles["Heading2"]

    skelp_ratio_img, skelp_amount_img = _skelp_impact_charts(latest_df, u)

    story = [
        Paragraph("Daily Repair Rate Report", title_style),
        Paragraph(
            f"Report Date: {pd.to_datetime(selected_date).strftime('%d.%m.%Y')}  |  "
            f"Active Projects: {latest_df['project_no'].nunique()}  |  "
            f"Coil and Plate Repair Rate: {current_ratio * 100:.2f}%",
            subtitle_style,
        ),
        Spacer(1, 0.4 * cm),
        _trend_chart(overall_ratio),
        Spacer(1, 0.3 * cm),
        _type_trend_chart(master_df, baseline_df),
        Spacer(1, 0.3 * cm),
        _worst_projects_chart(latest_df),
        Spacer(1, 0.3 * cm),
        _side_by_side(skelp_ratio_img, skelp_amount_img),
        Spacer(1, 0.3 * cm),
        _pareto_chart(
            latest_df, "repair_ratio", "Repair Ratio (%)", "Repair Ratio Pareto — Cumulative Contribution (Latest Day)",
            as_pct=True,
        ),
        Spacer(1, 0.3 * cm),
        _pareto_chart(
            latest_df,
            "total_repair_amount",
            f"Repair Amount ({u})",
            f"Repair Amount Pareto — Cumulative Contribution ({u}, Latest Day)",
            as_pct=False,
        ),
        Spacer(1, 0.3 * cm),
        _bubble_chart(
            latest_df,
            "repaired_spiral_length",
            f"Total Spiral Length ({u})",
            f"Repair Ratio vs. Production Volume by Project ({u}, Spiral Length, Latest Day)",
        ),
        Spacer(1, 0.3 * cm),
        _bubble_chart(
            latest_df,
            "project_total_pipe_length",
            f"Total Production Pipe Length ({u})",
            f"Repair Ratio vs. Production Volume by Project ({u}, Pipe Length, Latest Day)",
        ),
        Spacer(1, 0.3 * cm),
        _bubble_chart(
            _dimension_bubble_frame(master_df, selected_date, display_unit),
            "repaired_spiral_length",
            f"Total Spiral Length ({u})",
            f"Repair Ratio vs. Production Volume by Dimension ({u}, Spiral Length, Latest Day)",
            label_col="dimensions",
        ),
        Spacer(1, 0.3 * cm),
        _dimension_ratio_chart(master_df, selected_date),
        Spacer(1, 0.3 * cm),
        _amount_chart(daily_amount, u),
        Spacer(1, 0.5 * cm),
        Paragraph("Latest Day — Project Details", section_style),
        Spacer(1, 0.2 * cm),
        _detail_table(latest_df),
    ]
    doc.build(story)
    return buffer.getvalue()
