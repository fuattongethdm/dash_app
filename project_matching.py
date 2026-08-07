"""Best-guess matching between pipe_repair_details.project_sheet (a raw Excel
tab name) and repair_rates' (project_no, dimensions) pairs.

The two data sources are parsed independently and use different naming
conventions for the same underlying project (see project_sheet_links in
database.py) — this module only proposes candidates for a human to confirm,
it never writes anything itself.
"""

from __future__ import annotations

import re
from difflib import SequenceMatcher

import pandas as pd

_QUARTER_PREFIX_RE = re.compile(r"^\d{4}Q-", re.IGNORECASE)
_ID_RE = re.compile(r"(\d{1,2}-\d{2,3})")
_TRAILING_LETTERS_RE = re.compile(r"[A-Za-z]+$")
_PAREN_QUALIFIER_RE = re.compile(r"\s*\([^)]*\)\s*$")

_ID_SCORE_WEIGHT = 0.6
_DIM_SCORE_WEIGHT = 0.3
_SUFFIX_SCORE_WEIGHT = 0.1
_ID_FUZZY_CUTOFF = 0.6


def normalize_project_id(project_no: object) -> str:
    """'2026Q-10-108JT' -> '10-108'; '2025Q-05-093 (PT)' -> '05-093'."""
    text = _QUARTER_PREFIX_RE.sub("", str(project_no).strip())
    text = _PAREN_QUALIFIER_RE.sub("", text)
    text = _TRAILING_LETTERS_RE.sub("", text).strip()
    match = _ID_RE.search(text)
    return match.group(1) if match else text


def normalize_dimension_tokens(dimensions: object) -> list[str]:
    """'68"x.375"' -> ['68','375']; '54"x1.2"' -> ['54','12']."""
    tokens: list[str] = []
    for part in re.split(r"[xX]", str(dimensions)):
        digits = re.sub(r"[^0-9]", "", part).lstrip("0")
        if digits:
            tokens.append(digits)
    return tokens


def split_sheet_name(project_sheet: object) -> tuple[str, list[str], list[str]]:
    """'10-108 (68x375)' -> ('10-108', ['68','375'], [])
    '05-093 PT'        -> ('05-093', [], ['PT'])
    """
    text = str(project_sheet).strip()
    id_match = _ID_RE.match(text)
    id_part = id_match.group(1) if id_match else text
    remainder = text[id_match.end():] if id_match else ""

    dim_tokens: list[str] = []
    paren_match = re.search(r"\(([^)]*)\)", remainder)
    if paren_match:
        dim_tokens = normalize_dimension_tokens(paren_match.group(1))
        remainder = remainder[: paren_match.start()] + remainder[paren_match.end():]

    suffix_words = re.findall(r"[A-Za-z]+\d*", remainder)
    return id_part, dim_tokens, suffix_words


def build_candidate_pool(master_df: pd.DataFrame) -> pd.DataFrame:
    """Distinct (project_no, dimensions) pairs from repair_rates, pre-normalized."""
    columns = ["project_no", "dimensions", "production_type", "norm_id", "dim_tokens"]
    if master_df.empty:
        return pd.DataFrame(columns=columns)

    pool = master_df[["project_no", "dimensions", "production_type"]].drop_duplicates().reset_index(drop=True)
    pool["norm_id"] = pool["project_no"].map(normalize_project_id)
    pool["dim_tokens"] = pool["dimensions"].map(normalize_dimension_tokens)
    return pool


def score_candidate(sheet_id: str, sheet_dims: list[str], sheet_suffixes: list[str], candidate: dict) -> float:
    if sheet_id == candidate["norm_id"]:
        id_score = 1.0
    else:
        id_score = SequenceMatcher(None, sheet_id, candidate["norm_id"]).ratio()
    if id_score < _ID_FUZZY_CUTOFF:
        return round(id_score * 0.3, 4)

    if sheet_dims and candidate["dim_tokens"]:
        set_a, set_b = set(sheet_dims), set(candidate["dim_tokens"])
        dim_score = len(set_a & set_b) / len(set_a | set_b)
    else:
        dim_score = 0.5  # neutral -- nothing comparable on one or both sides

    suffix_bonus = 0.0
    if sheet_suffixes:
        suffix_text = " ".join(sheet_suffixes).lower()
        candidate_text = f"{candidate['project_no']} {candidate['production_type']}".lower()
        suffix_bonus = _SUFFIX_SCORE_WEIGHT * SequenceMatcher(None, suffix_text, candidate_text).ratio()

    return round(_ID_SCORE_WEIGHT * id_score + _DIM_SCORE_WEIGHT * dim_score + suffix_bonus, 4)


def suggest_match_for_sheet(project_sheet: str, candidate_pool: pd.DataFrame) -> list[dict]:
    """Ranked candidates best-first: [{'project_no', 'dimensions', 'score'}, ...]."""
    sheet_id, sheet_dims, sheet_suffixes = split_sheet_name(project_sheet)
    scored = [
        {
            "project_no": row["project_no"],
            "dimensions": row["dimensions"],
            "score": score_candidate(sheet_id, sheet_dims, sheet_suffixes, row),
        }
        for _, row in candidate_pool.iterrows()
    ]
    return sorted(scored, key=lambda item: item["score"], reverse=True)
