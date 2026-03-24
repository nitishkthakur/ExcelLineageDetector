"""Core data models for the Business Contract."""
from __future__ import annotations

import hashlib
from dataclasses import dataclass, field


def _make_id(*parts: str) -> str:
    raw = "|".join(parts)
    return hashlib.sha256(raw.encode()).hexdigest()[:12]


@dataclass
class ContractVariable:
    """A named business quantity -- a vector or scalar."""

    id: str
    business_name: str                    # LLM-inferred
    excel_location: str                   # "Sheet1!B3:B15"
    sheet: str
    cell_range: str
    direction: str                        # "row" | "column" | "scalar"
    length: int
    sample_values: list[float] = field(default_factory=list)
    variable_type: str = "intermediate"   # "input" | "output" | "intermediate"
    source_type: str = "formula"          # "hardcoded" | "formula" | "connection" | "external_link"
    upstream_source: str | None = None    # "upstream_a.xlsx!Data!A1:A10"
    upstream_file: str | None = None
    upstream_sheet: str | None = None
    upstream_range: str | None = None
    confidence: float = 0.0
    match_type: str | None = None        # "exact" | "approximate" | None

    @classmethod
    def from_vector(
        cls,
        sheet: str,
        cell_range: str,
        direction: str,
        length: int,
        values: list[float],
        source_type: str = "hardcoded",
    ) -> ContractVariable:
        vid = _make_id(sheet, cell_range)
        location = f"{sheet}!{cell_range}"
        return cls(
            id=vid,
            business_name="",  # filled by LLM later
            excel_location=location,
            sheet=sheet,
            cell_range=cell_range,
            direction=direction,
            length=length,
            sample_values=values[:5],
            source_type=source_type,
        )


@dataclass
class TransformationStep:
    """One formula transformation in a chain."""

    id: str
    output_variable_id: str
    input_variable_ids: list[str] = field(default_factory=list)
    excel_formula: str = ""
    sql_formula: str = ""                 # converted SQL-like notation
    sheet: str = ""
    cell_range: str = ""

    @classmethod
    def make(
        cls,
        output_id: str,
        input_ids: list[str],
        excel_formula: str,
        sql_formula: str,
        sheet: str,
        cell_range: str,
    ) -> TransformationStep:
        tid = _make_id("tx", output_id, excel_formula[:50])
        return cls(
            id=tid,
            output_variable_id=output_id,
            input_variable_ids=input_ids,
            excel_formula=excel_formula,
            sql_formula=sql_formula,
            sheet=sheet,
            cell_range=cell_range,
        )


@dataclass
class DependencyEdge:
    """Edge in the dependency graph."""

    source_id: str
    target_id: str
    edge_type: str    # "formula" | "external_link" | "vector_match" | "connection"
    metadata: dict = field(default_factory=dict)


@dataclass
class BusinessContract:
    """The complete contract output."""

    model_file: str
    output_sheets: list[str]
    variables: list[ContractVariable] = field(default_factory=list)
    transformations: list[TransformationStep] = field(default_factory=list)
    edges: list[DependencyEdge] = field(default_factory=list)
    connections: list[dict] = field(default_factory=list)
    upstream_lineage: list[dict] = field(default_factory=list)
