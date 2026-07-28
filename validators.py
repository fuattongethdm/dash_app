from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Iterable

import pandas as pd


ALLOWED_STATUS_MAP = {
    "completed": "Completed",
    "complete": "Completed",
    "in progress": "In Progress",
    "in-progress": "In Progress",
    "inprogress": "In Progress",
}

REQUIRED_COLUMNS = [
    "date",
    "production_type",
    "project_no",
    "dimensions",
    "qty",
    "project_total_pipe_length",
    "repaired_pipes_total_length",
    "repaired_spiral_length",
    "total_repair_amount",
    "total_repair_amount_incl_skelp",
    "project_status",
    "repair_ratio",
    "repair_ratio_incl_skelp",
]

NUMERIC_COLUMNS = [
    "qty",
    "project_total_pipe_length",
    "repaired_pipes_total_length",
    "repaired_spiral_length",
    "total_repair_amount",
    "total_repair_amount_incl_skelp",
    "repair_ratio",
    "repair_ratio_incl_skelp",
]


@dataclass
class ValidationReport:
    checks: dict[str, bool] = field(default_factory=dict)
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    import_rows: int = 0
    update_rows: int = 0
    insert_rows: int = 0

    @property
    def ok(self) -> bool:
        return not self.errors and all(self.checks.values())

    def add_check(self, name: str, ok: bool) -> None:
        self.checks[name] = ok

    def add_error(self, message: str) -> None:
        self.errors.append(message)


def normalize_spaces(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_status(value: object) -> str:
    cleaned = normalize_spaces(value)
    key = cleaned.lower()
    return ALLOWED_STATUS_MAP.get(key, cleaned.title() if cleaned else "")


def coerce_report_date(value: object) -> date | None:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return parsed.date()


def coerce_number(value: object) -> float | None:
    if value is None or value == "":
        return None
    if isinstance(value, str):
        value = value.replace(",", "").strip()
    parsed = pd.to_numeric(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return float(parsed)


def validate_dataframe(df: pd.DataFrame, report: ValidationReport) -> ValidationReport:
    report.add_check("Main table rows found?", not df.empty)
    row_count_ok = 1 <= len(df) <= 35
    report.add_check("Row count reasonable?", row_count_ok)
    if not row_count_ok:
        report.add_error(f"Row count outside the expected range (1-35): {len(df)}")
    report.import_rows = len(df)

    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    report.add_check("Required columns present?", not missing)
    if missing:
        report.add_error(f"Missing parser columns: {', '.join(missing)}")
        return report

    for idx, row in df.iterrows():
        excel_row = int(row.get("excel_row", idx + 4))
        for col in ["date", "production_type", "project_no", "dimensions"]:
            if pd.isna(row[col]) or str(row[col]).strip() == "":
                report.add_error(f"Excel row {excel_row}: required field empty: {col}")

        for col in NUMERIC_COLUMNS:
            if pd.isna(row[col]):
                report.add_error(f"Excel row {excel_row}: not numeric or empty: {col}")

        rr = row.get("repair_ratio")
        rr_i = row.get("repair_ratio_incl_skelp")
        amt = row.get("total_repair_amount")
        amt_i = row.get("total_repair_amount_incl_skelp")

        if pd.notna(rr) and not 0 <= rr <= 1:
            report.add_error(f"Excel row {excel_row}: Repair Ratio outside 0-1: {rr}")
        if pd.notna(rr_i) and not 0 <= rr_i <= 1:
            report.add_error(f"Excel row {excel_row}: Repair Ratio incl. Skelp outside 0-1: {rr_i}")
        if pd.notna(rr) and pd.notna(rr_i) and rr_i < rr:
            report.add_error(f"Excel row {excel_row}: Repair Ratio incl. Skelp is smaller than Repair Ratio")
        if pd.notna(amt) and pd.notna(amt_i) and amt_i < amt:
            report.add_error(f"Excel row {excel_row}: Total Repair Amount incl. Skelp is smaller than the normal value")

    report.add_check("Numeric columns are numeric?", not any("not numeric" in e for e in report.errors))
    report.add_check("Ratio values between 0-1?", not any("outside 0-1" in e for e in report.errors))
    report.add_check(
        "Incl. Skelp values smaller than normal values?",
        not any("incl. Skelp" in e and "smaller than" in e for e in report.errors),
    )
    return report


def mark_duplicate_counts(existing_keys: Iterable[tuple[str, str, str]], df: pd.DataFrame, report: ValidationReport) -> ValidationReport:
    existing = set(existing_keys)
    incoming_keys = list(zip(df["date"].astype(str), df["project_no"], df["dimensions"]))
    report.update_rows = sum(1 for key in incoming_keys if key in existing)
    report.insert_rows = len(incoming_keys) - report.update_rows
    return report
