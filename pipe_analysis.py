"""
Aggregation helpers for pipe-level repair data (see project_parser.py).
"""

from __future__ import annotations

import pandas as pd


def summarize_pipe_totals_by_sheet(pipe_df: pd.DataFrame) -> pd.DataFrame:
    """One row per project sheet: pipe count, total/avg/max repair ratio."""
    if pipe_df.empty:
        return pipe_df
    return (
        pipe_df.groupby("project_sheet", as_index=False)
        .agg(
            pipe_count=("pipe_no", "count"),
            total_repair_amount=("repair_amount", "sum"),
            avg_repair_ratio=("repair_ratio", "mean"),
            max_repair_ratio=("repair_ratio", "max"),
        )
        .sort_values("total_repair_amount", ascending=False)
    )


def worst_pipes(pipe_df: pd.DataFrame, top_n: int = 15) -> pd.DataFrame:
    """The top_n pipes with the highest repair ratio."""
    if pipe_df.empty:
        return pipe_df
    return pipe_df.sort_values("repair_ratio", ascending=False).head(top_n)
