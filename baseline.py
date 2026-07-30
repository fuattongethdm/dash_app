from __future__ import annotations

from typing import BinaryIO

import pandas as pd

from calculations import METERS_PER_FOOT


BASELINE_COLUMNS = [
    "project_no",
    "repaired_spiral_length",
    "total_repair_amount",
    "total_repair_amount_incl_skelp",
]

LEGACY_BASELINE_ASSIGNMENTS = {
    "05-092": {"reporting_year": 2025, "production_type": "Coil"},
    "12-112": {"reporting_year": 2026, "production_type": "Plate", "include_in_dashboard": False},
}


def _infer_production_type(project_no: object) -> str:
    """Same naming convention the daily-table parser uses (parser._production_type):
    project numbers for plate jobs carry "(PT)" or "Plate" in the name."""
    lowered = str(project_no).lower()
    if "(pt)" in lowered or "plate" in lowered:
        return "Plate"
    return "Coil"


def enrich_historical_baseline_metadata(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if out.empty:
        return out

    if "reporting_year" not in out.columns:
        out["reporting_year"] = pd.NA
    if "production_type" not in out.columns:
        out["production_type"] = pd.NA
    if "include_in_dashboard" not in out.columns:
        out["include_in_dashboard"] = True

    normalized_projects = out["project_no"].astype(str).str.strip()
    for project_no, assignment in LEGACY_BASELINE_ASSIGNMENTS.items():
        mask = normalized_projects.eq(project_no)
        out.loc[mask & out["reporting_year"].isna(), "reporting_year"] = assignment["reporting_year"]
        out.loc[mask & out["production_type"].isna(), "production_type"] = assignment["production_type"]
        if "include_in_dashboard" in assignment:
            out.loc[mask, "include_in_dashboard"] = assignment["include_in_dashboard"]

    out["reporting_year"] = pd.to_numeric(out["reporting_year"], errors="coerce").fillna(2026).astype(int)
    missing_type = out["production_type"].isna() | out["production_type"].astype(str).str.strip().eq("")
    out.loc[missing_type, "production_type"] = normalized_projects[missing_type].map(_infer_production_type)
    out["production_type"] = out["production_type"].astype(str).str.strip()
    out["include_in_dashboard"] = out["include_in_dashboard"].fillna(True).astype(bool)
    return out


def parse_historical_baseline_csv(file: BinaryIO) -> tuple[pd.DataFrame, list[str]]:
    df = pd.read_csv(file)
    df.columns = [str(col).strip().lower() for col in df.columns]
    missing = [col for col in BASELINE_COLUMNS if col not in df.columns]
    if missing:
        return pd.DataFrame(), [f"Historical baseline CSV eksik kolon iceriyor: {', '.join(missing)}"]

    if "dimensions" not in df.columns:
        df["dimensions"] = "Historical Baseline"
    if "project_status" not in df.columns:
        df["project_status"] = "Completed"
    if "reporting_year" not in df.columns:
        df["reporting_year"] = pd.NA
    if "production_type" not in df.columns:
        df["production_type"] = pd.NA

    df = df[
        [
            "project_no",
            "dimensions",
            "repaired_spiral_length",
            "total_repair_amount",
            "total_repair_amount_incl_skelp",
            "project_status",
            "reporting_year",
            "production_type",
        ]
    ].copy()
    df["project_no"] = df["project_no"].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    df["dimensions"] = df["dimensions"].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)
    df["project_status"] = df["project_status"].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)

    for col in ["repaired_spiral_length", "total_repair_amount", "total_repair_amount_incl_skelp"]:
        df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df[df["project_no"] != ""]
    errors: list[str] = []
    for idx, row in df.iterrows():
        row_no = idx + 2
        if pd.isna(row["repaired_spiral_length"]) or row["repaired_spiral_length"] <= 0:
            errors.append(f"CSV satir {row_no}: repaired_spiral_length pozitif numeric olmali.")
        if pd.isna(row["total_repair_amount"]) or row["total_repair_amount"] < 0:
            errors.append(f"CSV satir {row_no}: total_repair_amount numeric ve negatif olmayan deger olmali.")
        if pd.isna(row["total_repair_amount_incl_skelp"]) or row["total_repair_amount_incl_skelp"] < 0:
            errors.append(f"CSV satir {row_no}: total_repair_amount_incl_skelp numeric ve negatif olmayan deger olmali.")
        if (
            pd.notna(row["total_repair_amount"])
            and pd.notna(row["total_repair_amount_incl_skelp"])
            and row["total_repair_amount_incl_skelp"] < row["total_repair_amount"]
        ):
            errors.append(f"CSV satir {row_no}: incl. skelp amount normal amount degerinden kucuk.")

    if errors:
        return df, errors

    df = enrich_historical_baseline_metadata(df)
    denominator_m = df["repaired_spiral_length"] * METERS_PER_FOOT
    df["repair_ratio"] = df["total_repair_amount"] / denominator_m
    df["repair_ratio_incl_skelp"] = df["total_repair_amount_incl_skelp"] / denominator_m
    return df, []


def baseline_template_csv() -> str:
    return (
        "project_no,dimensions,repaired_spiral_length,total_repair_amount,total_repair_amount_incl_skelp,"
        "project_status,reporting_year,production_type\n"
        "05-092,Historical Baseline,26035.18,332.55,965.55,Completed,2025,Coil\n"
    )
