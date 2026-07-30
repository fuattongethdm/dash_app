"""
Links each daily-table row (project_no, dimensions) to its corresponding
per-project sheet tab (used by project_parser.py for pipe-level data).

The two live in the same workbook but aren't explicitly connected: the
daily table's project_no is free-text typed by hand (inconsistent — e.g.
"12-112 (54 x 1.2)" has dimensions baked into the project number, "JT"/
"SN"/etc. suffixes vary), while project sheet tabs follow their own naming
convention ("10-108 (68x375)", "05-093 PT", "11-111(90XX)").

Matching strategy (tiered, most-specific first): both sides are reduced to
a "core" project id (the leading NN-NNN number, zero-padding normalized)
plus a set of distinguishing tokens (diameter, thickness, and any text
suffix like "PT"/"HDD"/"#2"). A row matches a sheet when they share a core
and the tokens narrow it down to exactly one candidate. Ambiguous or
zero-candidate rows are left unmatched rather than guessed — silently
mislabeling pipe data to the wrong project is worse than leaving it blank.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import pandas as pd

_YEAR_PREFIX_RE = re.compile(r"^\d{4}Q-")
_CORE_RE = re.compile(r"^(\d{1,2}-\d{2,3})")
_TOKEN_RE = re.compile(r"[A-Za-z0-9\.]+")
_NUMBER_RE = re.compile(r"(\d+(?:\.\d+)?)")
_SUFFIX_CODES = {"jt", "sn", "nn", "rz", "ts", "je"}


def _core_id(text: object) -> str:
    text = _YEAR_PREFIX_RE.sub("", str(text)).strip()
    match = _CORE_RE.match(text)
    core = match.group(1) if match else text
    prefix, sep, number = core.partition("-")
    if sep and number.isdigit():
        return f"{prefix}-{int(number):03d}"
    return core


def _strip_core(text: object) -> str:
    text = _YEAR_PREFIX_RE.sub("", str(text)).strip()
    return _CORE_RE.sub("", text).strip()


def _tokens(text: object) -> set[str]:
    out: set[str] = set()
    for raw in _TOKEN_RE.findall(_strip_core(text)):
        token = raw.lower()
        if token in _SUFFIX_CODES:
            continue
        out.add(token)
        out.add(token.rstrip("."))  # "2." (from "(2.)") should match "2" (from "#2")
    return out


def _diameter_thickness(text: object, is_sheet: bool) -> tuple[str | None, str | None, bool]:
    rest = _strip_core(text) if is_sheet else str(text)
    is_combined = "combined" in rest.lower()
    if is_sheet:
        # A bare "xx" (no adjacent digit, e.g. "(84xx)") also marks the
        # sheet as the diameter-only/combined bucket for that diameter.
        is_combined = is_combined or bool(re.search(r"xx", rest, re.IGNORECASE) and not re.search(r"\d+x\d", rest))
    numbers = _NUMBER_RE.findall(rest)
    diameter = numbers[0] if numbers else None
    thickness = numbers[1] if len(numbers) > 1 else None
    return diameter, thickness, is_combined


@dataclass
class ProjectSheetMatchReport:
    matched_rows: int = 0
    unmatched_rows: int = 0
    unmatched: list[dict] = field(default_factory=list)


def match_project_sheets(main_df: pd.DataFrame, sheet_names: list[str]) -> tuple[pd.Series, ProjectSheetMatchReport]:
    """Return a Series aligned to main_df.index with the matched sheet name
    (or None), plus a report of what couldn't be matched."""
    report = ProjectSheetMatchReport()

    sheet_info = []
    for sheet in sheet_names:
        diameter, thickness, combined = _diameter_thickness(sheet, is_sheet=True)
        sheet_info.append(
            {
                "sheet": sheet,
                "core": _core_id(sheet),
                "tokens": _tokens(sheet),
                "diameter": diameter,
                "thickness": thickness,
                "combined": combined,
            }
        )

    results = []
    for _, row in main_df.iterrows():
        core = _core_id(row["project_no"])
        row_tokens = _tokens(row["project_no"]) | _tokens(row["dimensions"])
        diameter, thickness, combined = _diameter_thickness(row["dimensions"], is_sheet=False)

        same_core = [s for s in sheet_info if s["core"] == core]
        match = _resolve_candidate(same_core, row_tokens, diameter, thickness, combined)

        results.append(match)
        if match is None:
            report.unmatched_rows += 1
            report.unmatched.append({"project_no": row["project_no"], "dimensions": row["dimensions"]})
        else:
            report.matched_rows += 1

    return pd.Series(results, index=main_df.index, name="project_sheet"), report


def _resolve_candidate(
    same_core: list[dict],
    row_tokens: set[str],
    diameter: str | None,
    thickness: str | None,
    combined: bool,
) -> str | None:
    if len(same_core) == 1:
        return same_core[0]["sheet"]
    if not same_core:
        return None

    tiers = [
        lambda s: bool(s["tokens"] & row_tokens),
        lambda s: s["diameter"] == diameter and s["thickness"] == thickness,
        lambda s: s["diameter"] == diameter and s["combined"] == combined,
        lambda s: s["diameter"] == diameter,
    ]
    for tier in tiers:
        candidates = [s["sheet"] for s in same_core if tier(s)]
        if len(candidates) == 1:
            return candidates[0]
    return None
