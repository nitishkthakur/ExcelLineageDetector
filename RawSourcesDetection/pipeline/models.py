"""Data models for RawSourcesDetection."""
from __future__ import annotations

from dataclasses import dataclass, field


@dataclass
class SourceNode:
    """A file in the dependency graph (model or upstream)."""
    filename: str
    path: str               # full disk path if found; expected name if missing
    level: int              # 0 = model, 1 = direct upstream, 2 = upstream's upstream, …
    found_on_disk: bool
    sheets_referenced: list[str] = field(default_factory=list)


@dataclass
class FormulaRef:
    """One external formula reference found during tracing."""
    level: int              # at which level this ref was found
    source_file: str        # file containing the formula
    source_sheet: str
    source_cell: str
    target_file: str        # filename referenced in the formula
    target_sheet: str
    target_range: str
    file_found: bool        # whether target_file was found in inputs/
    resolved_path: str      # full disk path if found


@dataclass
class MissingFile:
    """A file referenced in formulas but not found in inputs/."""
    filename: str
    level: int              # tracing level at which it was first discovered
    referenced_by: str      # source file that references it
    sheets_needed: list[str] = field(default_factory=list)
    cells_referencing: list[str] = field(default_factory=list)  # "Sheet!Cell" strings


@dataclass
class MatchedVector:
    """A hardcoded model vector matched to an upstream source."""
    model_sheet: str
    model_range: str
    model_length: int
    model_sample: list[float]   # first 5 values
    match_type: str             # "exact" | "exact_subsequence" | "approximate"
    similarity: float           # 1.0 for exact matches
    upstream_file: str
    upstream_sheet: str
    upstream_range: str
    upstream_sample: list[float]


@dataclass
class UnmatchedVector:
    """A hardcoded model vector with no source match found."""
    model_sheet: str
    model_range: str
    model_length: int
    model_sample: list[float]


@dataclass
class RawSource:
    """A raw data source connection (ODBC, OLE DB, Power Query, Excel link, etc.)."""
    source_file: str    # which file this was detected in
    category: str       # database | file | powerquery | formula | vba | …
    sub_type: str       # odbc | oledb | sql_server | xlsx | csv | …
    connection: str     # raw connection string / path / URL
    location: str       # where in the file it was found (e.g. "Sheet1!A1")


@dataclass
class DetectionResult:
    """Complete output from the RawSourcesDetection pipeline."""
    model_file: str
    source_nodes: list[SourceNode]       # all files in the dependency tree
    formula_refs: list[FormulaRef]       # all formula-based external references
    missing_files: list[MissingFile]     # files referenced but not on disk
    matched_vectors: list[MatchedVector] # hardcoded vectors with source found
    unmatched_vectors: list[UnmatchedVector]  # hardcoded vectors with no source
    raw_sources: list[RawSource]         # ODBC/OLE DB/PQ connections across all files
