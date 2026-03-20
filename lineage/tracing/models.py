"""Data models for upstream tracing."""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class TracingVector:
    """A numeric vector with full values, used for upstream tracing."""
    file: str          # source file name
    sheet: str
    cell_range: str
    direction: str     # "column" or "row"
    length: int
    start_cell: str
    end_cell: str
    values: tuple[float, ...]  # ALL values (tuple for hashability)


@dataclass
class VectorMatch:
    """A match between a model vector and an upstream vector."""
    model_sheet: str
    model_range: str
    model_direction: str
    model_length: int
    model_sample: list[float]       # first 5 values

    match_rank: int                 # 1-based rank within this model vector's matches
    match_type: str                 # "exact" | "exact_subsequence" | "approximate"
    similarity: float               # 1.0 for exact, <1.0 for approximate

    upstream_file: str
    upstream_sheet: str
    upstream_range: str             # full upstream vector range
    upstream_direction: str
    upstream_length: int
    upstream_sample: list[float]    # first 5 values

    upstream_matched_range: str = ""  # specific sub-range that matched (for subsequences)
