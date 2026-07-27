from __future__ import annotations

import pandas as pd


METERS_PER_FOOT = 0.3048
FEET_PER_METER = 1 / METERS_PER_FOOT


def unit_label(display_unit: str) -> str:
    return "ft" if display_unit == "ft" else "m"


def amount_in_display_unit(values, display_unit: str):
    if display_unit == "ft":
        return values * FEET_PER_METER
    return values


def length_in_display_unit(values, display_unit: str):
    if display_unit == "m":
        return values * METERS_PER_FOOT
    return values


def normalize_repair_trend_project_key(project_no) -> str:
    value = str(project_no).strip()
    value = pd.Series([value]).str.replace(r"\s+", " ", regex=True).iloc[0]
    value = pd.Series([value]).str.replace(r"\s+#\d+$", "", regex=True).iloc[0]
    return pd.Series([value]).str.replace(r"\s+-\s+\d+$", "", regex=True).iloc[0]


def apply_meter_based_repair_ratios(df: pd.DataFrame) -> pd.DataFrame:
    """Return a copy with repair ratios recalculated as meter / meter.

    Excel length columns are in feet; repair amount columns are in meters.
    The dashboard standard is:
      repair amount (m) / repaired spiral length (ft * 0.3048)
    """
    out = df.copy()
    denominator_m = out["repaired_spiral_length"] * METERS_PER_FOOT
    denominator_m = denominator_m.where(denominator_m != 0)
    out["repair_ratio"] = out["total_repair_amount"] / denominator_m
    out["repair_ratio_incl_skelp"] = out["total_repair_amount_incl_skelp"] / denominator_m
    out[["repair_ratio", "repair_ratio_incl_skelp"]] = out[["repair_ratio", "repair_ratio_incl_skelp"]].fillna(0)
    return out


def daily_weighted_repair_ratios(df: pd.DataFrame, baseline_df: pd.DataFrame | None = None) -> pd.DataFrame:
    grouped = (
        df.groupby("date", as_index=False)
        .agg(
            total_repair_amount=("total_repair_amount", "sum"),
            total_repair_amount_incl_skelp=("total_repair_amount_incl_skelp", "sum"),
            repaired_spiral_length=("repaired_spiral_length", "sum"),
        )
        .sort_values("date")
    )
    baseline_repair = 0.0
    baseline_repair_incl = 0.0
    baseline_spiral = 0.0
    if baseline_df is not None and not baseline_df.empty:
        baseline_repair = baseline_df["total_repair_amount"].sum()
        baseline_repair_incl = baseline_df["total_repair_amount_incl_skelp"].sum()
        baseline_spiral = baseline_df["repaired_spiral_length"].sum()

    denominator_m = (grouped["repaired_spiral_length"] + baseline_spiral) * METERS_PER_FOOT
    denominator_m = denominator_m.where(denominator_m != 0)
    grouped["weighted_repair_ratio"] = ((grouped["total_repair_amount"] + baseline_repair) / denominator_m).fillna(0)
    grouped["weighted_repair_ratio_incl_skelp"] = (
        (grouped["total_repair_amount_incl_skelp"] + baseline_repair_incl) / denominator_m
    ).fillna(0)
    return grouped


def daily_weighted_repair_ratios_for_type(
    df: pd.DataFrame,
    production_type: str,
    baseline_df: pd.DataFrame | None = None,
) -> pd.DataFrame:
    data = df.copy()
    if "production_type" not in data.columns:
        data["production_type"] = "Coil"
    type_data = data[data["production_type"].astype(str).eq(str(production_type))].copy()

    type_baseline = pd.DataFrame()
    if baseline_df is not None and not baseline_df.empty:
        type_baseline = baseline_df.copy()
        if "production_type" not in type_baseline.columns:
            type_baseline["production_type"] = "Coil"
        type_baseline = type_baseline[type_baseline["production_type"].astype(str).eq(str(production_type))]

    return daily_weighted_repair_ratios(type_data, type_baseline)


def repair_amount_trend_data(df: pd.DataFrame, display_unit: str) -> pd.DataFrame:
    """Build repair amount trend data with non-negative daily increments.

    Daily repair is calculated by project/dimension against the previous report.
    Project revision suffixes like "#2" are treated as the same project so a
    rebaseline from 2025+2026 totals to 2026-only totals does not become a new
    daily repair amount.
    """
    data = df[["date", "project_no", "dimensions", "total_repair_amount"]].copy()
    data["date"] = pd.to_datetime(data["date"])
    data["repair_trend_project_key"] = data["project_no"].map(normalize_repair_trend_project_key)
    data["total_repair_amount_display"] = amount_in_display_unit(data["total_repair_amount"], display_unit)

    snapshot = (
        data.groupby("date", as_index=False)
        .agg(total_repair_amount_display=("total_repair_amount_display", "sum"))
        .sort_values("date")
    )
    if snapshot.empty:
        return snapshot.assign(daily_repair_amount_display=[])

    by_project = (
        data.groupby(["repair_trend_project_key", "dimensions", "date"], as_index=False)
        .agg(total_repair_amount_display=("total_repair_amount_display", "sum"))
        .sort_values(["repair_trend_project_key", "dimensions", "date"])
    )
    by_project["previous_total"] = by_project.groupby(["repair_trend_project_key", "dimensions"])[
        "total_repair_amount_display"
    ].shift(1)
    first_report_date = snapshot["date"].min()
    by_project["daily_delta"] = by_project["total_repair_amount_display"] - by_project["previous_total"]
    new_after_first_report = by_project["previous_total"].isna() & (by_project["date"] > first_report_date)
    by_project.loc[new_after_first_report, "daily_delta"] = by_project.loc[new_after_first_report, "total_repair_amount_display"]
    by_project["daily_delta"] = by_project["daily_delta"].clip(lower=0).fillna(0)

    daily = by_project.groupby("date", as_index=False).agg(daily_repair_amount_display=("daily_delta", "sum"))
    result = snapshot.merge(daily, on="date", how="left").fillna({"daily_repair_amount_display": 0})
    result["daily_repair_amount_display"] = result["daily_repair_amount_display"].clip(lower=0)
    return result
